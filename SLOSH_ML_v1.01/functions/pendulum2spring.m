function [Km,H] = pendulum2spring(Mm,Ln,g,Hi)
% PENDULUM2SPRING Converts pendulum parameters to spring parameters

% Spring constant
Km = (Mm .* g) ./ Ln;
%Spring hinge point 
H = Hi-Ln;

end

